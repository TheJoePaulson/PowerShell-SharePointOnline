<#

SHAREPOINT ONLINE ADMIN TOOLS

       _____ _                    _____      _       _         
      / ____| |                  |  __ \    (_)     | |       
     | (___ | |__   __ _ _ __ ___| |__) |__  _ _ __ | |_   
      \___ \| '_ \ / _' | '__/ _ \  ___/ _ \| | '_ \| __| 
      ____) | | | | (_| | | |  __/ |  | (_) | | | | | |_  
     |_____/|_| |_|\__,_|_|  \___|_|   \___/|_|_| |_|\__| 

               _           _         _______          _     
      /\      | |         (_)       |__   __|        | |    
     /  \   __| |_ __ ___  _ _ __      | | ___   ___ | |___ 
    / /\ \ / _' | '_ ' _ \| | '_ \     | |/ _ \ / _ \| / __|
   / ____ \ (_| | | | | | | | | | |    | | (_) | (_) | \__ \
  /_/    \_\__,_|_| |_| |_|_|_| |_|    |_|\___/ \___/|_|___/
                       
.SYNOPSIS
This script allows for the execution of all components of a complex SharePoint build, content migration,
or various other M365/SharePoint Online administrative functions.


REVISED:  3/18/2022

This is a culmination of scripts, processes, functions modified by the help of others and other resources. 

#>

## START UP MESSAGE
cls
Write-Host "Admin Tool is starting up, and loading dependencies..." -f Yellow

## Sets PowerShell to use HTTPS connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

## Get Execution Policy to setback at end of script
$varExecPolicy = Get-ExecutionPolicy

## Change Execution Policy for Script run
Set-ExecutionPolicy -ExecutionPolicy ByPass -Force

### REVISION DATE -- WILL BE DISPLAYED IN FORM!
$RevDate = "3/18/2022"

## Log ON ($true) or OFF ($false)
$log = $true

###################################################################
#            SCRIPT FUNCTIONS                                     #
###################################################################

## Static Variables
$ColumnName1 = "Created"
$ColumnName2 = "Author"
$ColumnName3 = "Modified"
$ColumnName4 = "Editor"
$ColumnName5 = "Title"
$ColumnName6 = "CheckoutUser"

## FUNCTION: Time Stamp
function Get-TimeStamp {
    
    return "[ {0:MM/dd/yy} {0:HH:mm:ss} ]" -f (Get-Date)
    
}

## Set Path

## Check for Modules/Cmdlets

## Logging

## End Script Function - Prereqs NOT Found
function PrereqsNotFound {
    $NoPrereqa = "*** PowerShell Prerequisites are not found on this machine.  Install Prerequisites first. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($NoPrereqa)" | Add-Content $logFile }
    Write-Host " "
    Write-Host "   *** " -f Red -NoNewline
    Write-Host "PowerShell Prerequisites are not found on this machine.  Install Prerequisites first." -f Yellow -NoNewline
    Write-Host " ***" -f Red -NoNewLine
    Write-Host "   "$(Get-TimeStamp)
    Write-Host " "
    Write-Host "Hit 'ENTER' KEY to exit."
    Write-Host " "
    pause

}

## Check if 'PnP' Module is installed correctly, install & import if module NOT found
Function CheckModPnP {
if (Get-Module -ListAvailable -Name "PnP.PowerShell") {
  $CheckModPnPa = "-- 'SP PnP' Module Installed and Ready to Go - $(Get-TimeStamp)"
  if ($log -eq $true) { "`n$($CheckModPnPa)" | Add-Content $logFile }
  Write-Host $CheckModPnPa

}
else {
    ## Prereqs Not Found
    PrereqsNotFound
  }
}

## Function to download a list template from SharePoint Online using powershell
Function Download-SPOListTemplate
{
    param
    (
        [string]$SiteURL  = $(throw "Enter the Site URL!"),
        [string]$ListTemplateName = $(throw "Enter the List Template Name!"),
        [string]$ExportFile = $(throw "Enter the File Name to Export List Template!")
    )
    Try {
        #Get Credentials to connect
        #$Cred= Get-Credential
        #$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($TenantAdmin, $TenantAdminPw)
  
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
         
        #Get the "List Templates" Library
        $List= $Ctx.web.Lists.GetByTitle("List Template Gallery")
        $ListTemplates = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
        $Ctx.Load($ListTemplates)
        $Ctx.ExecuteQuery()
 
        #Filter and get given List Template
        $ListTemplate = $ListTemplates | where { $_["TemplateTitle"] -eq $ListTemplateName }
 
        If($ListTemplate -ne $Null)
        {
            #Get the File from the List item
            $Ctx.Load($ListTemplate.File)
            $Ctx.ExecuteQuery()
 
            #Download the list template
            $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$ListTemplate.File.ServerRelativeUrl)
            $WriteStream = [System.IO.File]::Open($ExportFile,[System.IO.FileMode]::Create)
            $FileInfo.Stream.CopyTo($WriteStream)
            $WriteStream.Close()
 
            write-host -f Green "List Template Downloaded to $ExportFile!" $_.Exception.Message
        }
        else
        {
            Write-host -f Yellow "List Template Not Found:"$ListTemplateName
        }
    }
    Catch {
        write-host -f Red "Error Downloading List Template!" $_.Exception.Message
    }
}

## FUNCTION: Build Text Interface Layout
function MainHeaderLayout {
Write-Host "Rev. " -f Green -NoNewline
Write-Host $RevDate
Write-Host " ________________________________________________________________ "
Write-Host "|" -NoNewline
Write-Host "                                                                " -f Green -NoNewline
Write-Host "|" -f White

Write-Host "|" -NoNewline
Write-Host "       _____ _                    _____      _       _          " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "      / ____| |                  |  __ \    (_)     | |         " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "     | (___ | |__   __ _ _ __ ___| |__) |__  _ _ __ | |_        " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "      \___ \| '_ \ / _' | '__/ _ \  ___/ _ \| | '_ \| __|       " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "      ____) | | | | (_| | | |  __/ |  | (_) | | | | | |_        " -f Green -NoNewline
Write-Host "|"
Write-Host "|" -NoNewline
Write-Host "     |_____/|_| |_|\__,_|_|  \___|_|   \___/|_|_| |_|\__|       " -f Green -NoNewline
Write-Host "|"

Write-Host "|                                                                |"
Write-Host "|" -NoNewline
Write-Host "               _           _         _______          _         " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "      /\      | |         (_)       |__   __|        | |        " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "     /  \   __| |_ __ ___  _ _ __      | | ___   ___ | |___     " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "    / /\ \ / _' | '_ ' _ \| | '_ \     | |/ _ \ / _ \| / __|    " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "   / ____ \ (_| | | | | | | | | | |    | | (_) | (_) | \__ \    " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "  /_/    \_\__,_|_| |_| |_|_|_| |_|    |_|\___/ \___/|_|___/    " -f Green -NoNewline
Write-Host "|"

Write-Host "|" -NoNewline
Write-Host "                                                                " -f Green -NoNewline
Write-Host "|" -f White
Write-Host "|" -NoNewline
Write-Host "                                                                " -f Green -NoNewline
Write-Host "|" -f White

Write-Host "|                                                                |"
Write-Host "|   [ "  -NoNewline 

Write-Host "Tools and Functions for SharePoint Online Maintenance" -f Yellow -NoNewline
Write-Host " ]    |" -f White
Write-Host "|                                                                |" -f White
Write-Host "|________________________________________________________________|" -f White
Write-Host " "
Write-Host " "

}

## FUNCTION: Build Menu
Function MainMenuMaker{
    param(
        [parameter(Mandatory=$true)][String[]]$Selections,
        [switch]$IncludeExit,
        [string]$Title = $null
        )

    $Width = if($Title){$Length = $Title.Length;$Length2 = $Selections|%{$_.length}|Sort -Descending|Select -First 1;$Length2,$Length|Sort -Descending|Select -First 1}else{$Selections|%{$_.length}|Sort -Descending|Select -First 1}
    $Buffer = if(($Width*1.5) -gt 78){[math]::floor((78-$width)/2)}else{[math]::floor($width/4)}
    if($Buffer -gt 6){$Buffer = 6}
    $MaxWidth = $Buffer*2+$Width+$($Selections.count).length+2
    $Menu = @()
    $Menu += "╔"+"═"*$maxwidth+"╗"
    if($Title){
        $Menu += "║"+" "*[Math]::Floor(($maxwidth-$title.Length)/2)+$Title+" "*[Math]::Ceiling(($maxwidth-$title.Length)/2)+"║"
        $Menu += "╟"+"─"*$maxwidth+"╢"
    }
    For($i=1;$i -le $Selections.count;$i++){
        $Item = "$(if ($Selections.count -gt 9 -and $i -lt 10){" "})$i`. "
        $Menu += "║"+" "*$Buffer+$Item+$Selections[$i-1]+" "*($MaxWidth-$Buffer-$Item.Length-$Selections[$i-1].Length)+"║"
    }
    If($IncludeExit){
        $Menu += "║"+" "*$MaxWidth+"║"
        $Menu += "║"+" "*$Buffer+"X - Exit"+" "*($MaxWidth-$Buffer-8)+"║"
    }
    $Menu += "╚"+"═"*$maxwidth+"╝"
    $menu
}

## FUNCTION:  Call Menu to redraw MainMenuMaker -Selections 'Allow PnP Auth','Site Collection Build','Sub Site Build','Library Build/Index','Pre-Provision OneDrive','LISTS - Save as Templates','LISTS - Download Templates','LISTS - Upload and Create','Assign Permissions','EXPORT Navigation','IMPORT Navigation','UTIL:Add Admin to ALL Sites','UTIL:Remove Admin from ALL Sites' -Title 'Choose Migration Operation' -IncludeExit
function CallMenu {
    cls
    MainMenuMaker -Selections 'Allow PnP Auth','Site Collection Build','Sub Site Build','Library Build/Index','Pre-Provision OneDrive','Allow Scripts on Site','LISTS - Save as Templates','LISTS - Download Templates','LISTS - Upload and Create','Assign Permissions','EXPORT Navigation','IMPORT Navigation','UTIL:Add Admin to ALL Sites','UTIL:Remove Admin from ALL Sites' -Title 'Choose Migration Operation' -IncludeExit
}

## Connections:  PnP Cmdlets
function CmdPnpConnect {
## Connect to SP PnP
Try {
    Connect-PnPOnline -Url $tenantAdminUrl -Credential $TenantCredential
    $connectSPpnpa = "SPO PnP Service: CONNECTED - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($connectSPpnpa)" | Add-Content $logFile }
    Write-Host $connectSPpnpa -f Green
}
Catch {

    $connectSPpnpb = "    ## Error Connecting to SPO PnP service -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($connectSPpnpb)" | Add-Content $logFile }
    write-host $connectSPpnpb -f Red
}
}

## EXPORT SITE NAVIGATION SUB FUNCTION
function ExportSiteNavigation {

    Param (
    [parameter(Mandatory = $False)]
    [String]$SourceSite,

    [parameter(Mandatory = $False)]
    [String]$BackupDestination,

    [parameter(Mandatory = $False)]
    [String]$DestinationSite,

    [Parameter(Mandatory=$true)]
    [ValidateSet("TopNavigationBar", "Footer", "QuickLaunch", "SearchNav")]
    [String]$NavigationLocation

    )
    
    ## Build Credential to Connect
    $SrcNavTenantCreds = New-Object System.Management.Automation.PSCredential($SrcNavTenantAdmin,$SrcNavTenantAdminPw)

    Write-Host " "
    $ExportNav4 = " -- Connecting to the Source Site... $SrcNavUrl  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExportNav4)" | Add-Content $logFile }
    Write-Host $ExportNav4
    
    Connect-PnPOnline -Url $SrcNavUrl -Credentials $SrcNavTenantCreds


    #$FileDestination = "$BackupDestination"+"\SiteNavigationBackup.xlsx"
    ## Set file for export results
    #$NavExportFile = "$TempDir\SiteNavigationBackup.xlsx"
    $NavExportFile = ".\SiteNavigationBackup.xlsx"

    
    $MainNavigationData = Get-PnPNavigationNode  -Location $NavigationLocation

    Write-Host " "
    $ExportNav5 = " -- Collecting site navigation information...  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExportNav5)" | Add-Content $logFile }
    Write-Host $ExportNav5

foreach ($MainMenu in $MainNavigationData) {

    $ManinMenuID = $MainMenu.Id
    $MainMenuTitle = $MainMenu.Title 
    $MainMenuURL = $MainMenu.Url

    $FirstSubMenu = Get-PnPNavigationNode -Id  $ManinMenuID
    $FirstSubMenuChilder = $FirstSubMenu.Children

    $Export = [pscustomobject]@{
        MenuID       	    = $ManinMenuID
        MenuTitle			= $MainMenuTitle
        MenuURL     	    = $MainMenuURL
        MenuParent          = ""
        MenuPOS 			= "MainMenu"
        
        
    }
    $Export | Export-Excel -Path $NavExportFile -Append

    foreach ($FirstSubMenuChild in $FirstSubMenuChilder) {

        $FirstSubMenuID = $FirstSubMenuChild.Id
        $FirstSubMenuTitle = $FirstSubMenuChild.Title 
        $FirstSubMenuURL = $FirstSubMenuChild.Url

        $SecondSubMenu = Get-PnPNavigationNode -Id $FirstSubMenuID
        $SecondSubMenuChilder = $SecondSubMenu.Children

        $Export = [pscustomobject]@{
            MenuID       	    = $FirstSubMenuID
            MenuTitle			= $FirstSubMenuTitle
            MenuURL     	    = $FirstSubMenuURL
            MenuParent          = $MainMenu.Id
            MenuPOS 			= "FirstSubMenu"
            
            
        }
        $Export | Export-Excel -Path $NavExportFile -Append

        foreach ($SecondSubMenuChild in $SecondSubMenuChilder) {

            $SecondSubMenuID = $SecondSubMenuChild.Id 
            $SecondSubMenuTitle = $SecondSubMenuChild.Title
            $SecondSubMenuURL = $SecondSubMenuChild.Url

           $Export = [pscustomobject]@{
                MenuID       	    = $SecondSubMenuID
                MenuTitle			= $SecondSubMenuTitle
                MenuURL     	    = $SecondSubMenuURL
                MenuParent          = $FirstSubMenuID
                MenuPOS 			= "SecondSubMenu"
                
                
            }
            $Export | Export-Excel -Path $NavExportFile -Append
    
            
        }
        
    }

    
}
    Write-Host " "
    $ExportNav6 = " -- Site navigation EXPORT process is complete for: $SrcNavLocation on $SrcNavUrl  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExportNav6)" | Add-Content $logFile }
    Write-Host $ExportNav6 -f Green
   }

## IMPORT SITE NAVIGATION SUB FUNCTION
function ImportSiteNavigation {

    Param (
    
    [parameter(Mandatory = $False)]
    [String]$SourceSite,

    [parameter(Mandatory = $False)]
    [String]$BackupDestination,

    [parameter(Mandatory = $False)]
    [String]$DestinationSite,

    [Parameter(Mandatory=$true)]
    [ValidateSet("TopNavigationBar", "Footer", "QuickLaunch", "SearchNav")]
    [String]$NavigationLocation

    )

    Write-Host " "
    $ImportNav4 = " -- Connecting to the Destination Site... $DestNavUrl  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav4)" | Add-Content $logFile }
    Write-Host $ImportNav4
    
    Connect-PnPOnline -Url $DestNavUrl -Credentials $TenantCredential

    # Remove existing hub navigation
    Remove-PnPNavigationNode -All -Force
    sleep -Seconds 3

    ## DEBUG
    #pause

    Write-Host " "
    $ImportNav5 = " -- Removed existing site navigation elements.  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav5)" | Add-Content $logFile }
    Write-Host $ImportNav5


    $ExcelBackupFile = Import-Excel -Path $NavExportFile

    #Main Menu Scope
    ##############################################################################
    $MainMenuScope = $ExcelBackupFile | Where-Object {$_.MenuPOS -eq "MainMenu"}
    #$Path= $env:TEMP
    $Path = $TempDir

        Write-Host " "
        $ImportNav6 = " -- Restoring navigation elements to the destination site.  - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($ImportNav6)" | Add-Content $logFile }
        Write-Host $ImportNav6

    foreach ($MainMenu in $MainMenuScope) {

         $MenuTitle = $MainMenu.MenuTitle
         $MenuUrl = $MainMenu.MenuURL
         $MenuID = $MainMenu.MenuID

         $MainMenu = Add-PnPNavigationNode -Location $NavigationLocation -Title $MenuTitle -Url $MenuUrl -External
         $MainMenuNewId = $MainMenu.Id

         $TempExport = [pscustomobject]@{
            MenuID          	    = $MainMenuNewId
            MenuTitle		    	= $MenuTitle
            OldID                   = $MenuID
        
        }
        $TempExport | Export-Excel -Path "$Path\TempExport.xlsx" -Append

    
    }

    Write-Host " "
    $ImportNav7 = " -- Main-Menu elements restored successfully!  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav7)" | Add-Content $logFile }
    Write-Host $ImportNav7 -f Green

    ##############################################################################

    #First Sub-Menu Scope

    ##############################################################################
    $FirstSubMenuScope = $ExcelBackupFile | Where-Object {$_.MenuPOS -eq "FirstSubMenu"}
    $ImportTempScop = Import-Excel -Path "$Path\TempExport.xlsx"

    foreach ($FirstSubMenu in $FirstSubMenuScope) {

        $FirstSubTitle = $FirstSubMenu.MenuTitle
        $FirstSubUrl = $FirstSubMenu.MenuURL
        $FirstSubID = $FirstSubMenu.MenuID 
        $FirstSubOldParentId = $FirstSubMenu.MenuParent

        foreach ($Temp in $ImportTempScop) {
            $TempOldID = $Temp.OldID

            If ($FirstSubOldParentId -eq $TempOldID) {

                $FirstSub = Add-PnPNavigationNode -Location $NavigationLocation -Title $FirstSubTitle -Url $FirstSubUrl -Parent $Temp.MenuID -External
                $FirstSubNewId = $FirstSub.Id

                $TempExport = [pscustomobject]@{
                    MenuID          	    = $FirstSubNewId
                    MenuTitle		    	= $FirstSubTitle 
                    OldID                   = $FirstSubID
                
                }
                $TempExport | Export-Excel -Path "$Path\TempExport.xlsx" -Append
            }

        
        }
    
    }

    Write-Host " "
    $ImportNav8 = " -- First Sub-Menu elements restored successfully!  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav8)" | Add-Content $logFile }
    Write-Host $ImportNav8 -f Green

    ##############################################################################

    #Second Sub-Menu Scope

    ##############################################################################

    $SecondSubMenuScope = $ExcelBackupFile | Where-Object {$_.MenuPOS -eq "SecondSubMenu"}
    $ImportTempScop = Import-Excel -Path "$Path\TempExport.xlsx"

    foreach ($SecondSubMenu in $SecondSubMenuScope) {

        $SecondSubTitle = $SecondSubMenu.MenuTitle
        $SecondSubUrl = $SecondSubMenu.MenuURL
        $SecondSubID = $SecondSubMenu.MenuID 
        $SecondSubOldParentId = $SecondSubMenu.MenuParent

        foreach ($Temp in $ImportTempScop) {
            $TempOldID = $Temp.OldID

            If ($SecondSubOldParentId -eq $TempOldID) {

                $SecondSub = Add-PnPNavigationNode -Location $NavigationLocation -Title $SecondSubTitle -Url $SecondSubUrl -Parent $Temp.MenuID -External
                $FirstSubNewId = $FirstSub.Id

                $TempExport = [pscustomobject]@{
                    MenuID          	    = $SecondSubNewId
                    MenuTitle		    	= $SecondSubTitle 
                    OldID                   = $SecondSubID
                
                }
                $TempExport | Export-Excel -Path "$Path\TempExport.xlsx" -Append
            }

        
        }
    
    }

    Write-Host " "
    $ImportNav9 = " -- Second Sub-Menu elements restored successfully!  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav9)" | Add-Content $logFile }
    Write-Host $ImportNav9 -f Green

    Remove-Item -Path "$Path\TempExport.xlsx" -Recurse
    ## Disconnect SPO Connection
    Disconnect-PnPOnline

    Write-Host " "
    $ImportNav10 = " -- Site navigation IMPORT process is complete for: $DestNavLocation on $DestNavUrl  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav10)" | Add-Content $logFile }
    Write-Host $ImportNav10 -f Green

    }

## Create functions for EACH operation that can be called in menu as options

## 1) PnP Auth
function CmdPnPAuth {
Try {
    $PnPAuth1 = "You are about to AUTHORIZE PnP Cmdlets for this tenant. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($PnPAuth1)" | Add-Content $logFile }
    Write-Host $PnPAuth1 -f Yellow
    Write-Host "Follow the prompts..."
    #Connect Function
    Connect-PnPOnline -Url $tenantAdminUrl -PnPManagementShell
    $PnPAuth2 = "PnP Cmdlets have been succesfully authorized for this tenant. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($PnPAuth2)" | Add-Content $logFile }
    Write-Host $PnPAuth2 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B2"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C2"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate1 = "PnP Auth has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate1)" | Add-Content $logFile }
    
    pause
    CallMenu
}
Catch {

    $PnPAuth3 = "    ## Error Authorizing PnP Cmdlets on tenant: $tenantName -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($PnPAuth3)" | Add-Content $logFile }
    write-host $PnPAuth3 -f Red
    pause
    CallMenu
}
}

## 2) Site Collection Build
function CmdscBuild {
Try {
    $CSVFileSC = "$TempDir\SiteCollections.csv"

    #Get the CSV file
    $CSVFile1 = Import-Csv $CSVFileSC

    ## Connect to Tenant
    CmdPnpConnect

    ## Begin Site Collection Build
    $SCmsg1 = "-- BEGIN SharePoint Site Collection Build - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg1)" | Add-Content $logFile }

Try {
    #Get Site Collections to create from a CSV file
    ##################$SiteCollections1 = Import-Csv -Path $CSVFile1
    $SiteCollections1 = $CSVFile1
 
    #Loop through csv and create site collections from each row
    ForEach ($Site1 in $SiteCollections1)
    {
        #Get Parameters from CSV
        $Url = $Site1.Url
        $Title = $Site1.Title
        $Owner = $Site1.Owner
        $Template = $Site1.Template
        $TimeZone = $Site1.TimeZone
 
        #Check if site exists already
        $SiteExists = Get-PnPTenantSite | Where {$_.Url -eq $URL} 
        If ($SiteExists -eq $null)
        {
            #Create site collection
            $SCmsg2 = "   -- Creating Site Collection: " + $Url + " - $(Get-TimeStamp)"
            if ($log -eq $true) { "`n$($SCmsg2)" | Add-Content $logFile }
            Write-host $SCmsg2 -f Yellow
            New-PnPTenantSite -Url $Url -Title $Title -Owner $Owner -Template $Template -TimeZone $TimeZone -RemoveDeletedSite -ErrorAction Stop        
            Write-host "`t Done!" -f Green
        }
        Else
        {
            $SCmsg3 = "   ## Site $($Url) exists already! - $(Get-TimeStamp)"
            if ($log -eq $true) { "`n$($SCmsg3)" | Add-Content $logFile }
            write-host $SCmsg3 -foregroundcolor Yellow
        }
    }
}
Catch {
    $SCmsg4 = "   ## `tError:" + $_.Exception.Message + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg4)" | Add-Content $logFile }
    write-host $SCmsg4 -f Red 
}


###################################################################
#                       INDEX DEFAULT DOCS LIBRARY                #
###################################################################

    Write-Host " "
    $SCmsg5 = "--  Begin Column Indexing of Default Site Libraries (Documents) - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg5)" | Add-Content $logFile }
    Write-Host $SCmsg5

    #Get the CSV file
    $CSVFile2 = Import-Csv $CSVFileSC

    #Read CSV file and create document library indexes
    ForEach($Line2 in $CSVFile2) {

    Try {

    ## Build URL
    $SiteName = $Line2.SiteLocation
    $SiteURL = "https://$tenantName.sharepoint.com$SiteName"
    Write-Host $SiteURL
    ## Set Library Name
    #$LibraryName = $Line.Library
    $LibraryName = "Documents"

    #Connect to PNP Online
    Connect-PnPOnline -Url $SiteURL -Credential $TenantCredential
    Write-Host " "
    Write-Host "PnP Service: CONNECTED" -ForegroundColor Green
    $SCmsg6 = "   --  Begin Column Indexing of Default Site Libraries (Documents): " + $SiteURL + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg6)" | Add-Content $logFile }
    Write-Host $SCmsg6
 
    #Get the Context
    $Context = Get-PnPContext

    #Get the Field from List
    $Field1 = Get-PnPField -List $LibraryName -Identity $ColumnName1
    $Field2 = Get-PnPField -List $LibraryName -Identity $ColumnName2
    $Field3 = Get-PnPField -List $LibraryName -Identity $ColumnName3
    $Field4 = Get-PnPField -List $LibraryName -Identity $ColumnName4
    $Field5 = Get-PnPField -List $LibraryName -Identity $ColumnName5
    $Field6 = Get-PnPField -List $LibraryName -Identity $ColumnName6
 
    #Set the Indexed Property of the Field
    Write-Host " "
    Write-Host " --- Setting index for Column: " $ColumnName1
    $Field1.Indexed = $True
    $Field1.Update() 
    Write-Host " "
    Write-Host " --- Setting index for Column: " $ColumnName2
    $Field2.Indexed = $True
    $Field2.Update() 
    Write-Host " "
    Write-Host " --- Setting index for Column: " $ColumnName3
    $Field3.Indexed = $True
    $Field3.Update() 
    Write-Host " "
    Write-Host " --- Setting index for Column: " $ColumnName4
    $Field4.Indexed = $True
    $Field4.Update() 
    Write-Host " "
    Write-Host " --- Setting index for Column: " $ColumnName5
    $Field5.Indexed = $True
    $Field5.Update() 
    Write-Host " "
    Write-Host " --- Setting index for Column: " $ColumnName6
    $Field6.Indexed = $True
    $Field6.Update() 

    $Context.ExecuteQuery()
    Write-Host " "
    $SCmsg7 = "   --  Columns Indexed: $($ColumnName1),  $($ColumnName2), $($ColumnName3), $($ColumnName4), $($ColumnName5), $($ColumnName6) - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg7)" | Add-Content $logFile }
    $SCmsg8 = "   --  Column Indexing Complete for 'Documents' Library on " + $SiteURL + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg8)" | Add-Content $logFile }
    Write-Host $SCmsg8 -ForegroundColor Green

    ## Disconnect from PnP Connection
    Disconnect-PnPOnline

    }
    
    Catch {
        $SCmsg9 = "   ##  Error Indexing Column! " + $_.Exception.Message + " - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($SCmsg9)" | Add-Content $logFile }
        write-host $SCmsg9 -f Red 
          }

}


    ## Log Complete
    $SCmsg10 = "*** Site Collection Build and Index Operation Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg10)" | Add-Content $logFile }
    Write-Host $SCmsg10 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B3"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C3"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate2 = "Site Collection Build and Index has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate2)" | Add-Content $logFile }

    pause
    CallMenu
}
Catch {

    $SCmsg11 = "    ## Error(s) Building Site Collections -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SCmsg11)" | Add-Content $logFile }
    write-host $SCmsg11 -f Red
    pause
    CallMenu
}
}

## 3) Subsite Build
function CmdSubBuild {

    $CSVFileSubs = "$TempDir\Subsites.csv"

    #Get the CSV file
    $CSVFile1 = Import-Csv $CSVFileSubs

    ## Begin Subsite Build
    $SubMsg1 = "-- BEGIN SharePoint Subsite Build - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg1)" | Add-Content $logFile }

#Read CSV file for each subsite
ForEach($Line1 in $CSVFile1)
{
 
Try{
   
    ## Build URL
    $SubsiteUrl1 = $Line1.SubsiteUrl
    $SiteName1 = $Line1.ParentUrl
    $SiteURL1 = "https://$tenantName.sharepoint.com$SiteName1"
    $FullSiteURL1 = "https://$tenantName.sharepoint.com$SiteName1" + "/" + $SubsiteUrl1
    $SubMsg2 = "   -- Creating Subsite: " + $FullSiteURL1 + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg2)" | Add-Content $logFile }
    Write-Host $SubMsg2 -f Yellow

    #Connect to Parent URL
    Connect-PnPOnline -Url $SiteURL1 -Credentials $TenantCredential
    Write-Host " "
    $SubMsg3 = "   -- PnP Service: CONNECTED TO " + $SiteURL1 + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg3)" | Add-Content $logFile }
    Write-Host $SubMsg3  -f Green
    
    #sharepoint online powershell create subsite
    New-PnPWeb -Title $Line1.SubsiteTitle -Url $Line1.SubsiteUrl -Locale $Line1.Locale -Template $Line1.Template -ErrorAction Stop
    $SubMsg4 = "   -- Site Created Successfully! - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg4)" | Add-Content $logFile }
    Write-host $SubMsg4 -f Green

    ## Pause Briefly
    Start-Sleep -Seconds 5

}
Catch {
    $SubMsg5 = "   ## Error: " + $_.Exception.Message + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg5)" | Add-Content $logFile }
    write-host $SubMsg5 -f Red
}

}

###################################################################
#                       INDEX DEFAULT DOCS LIBRARY                #
###################################################################

Write-Host " "
$SubMsg6 = "--  Begin Column Indexing of Default Site Libraries (Documents) - $(Get-TimeStamp)"
if ($log -eq $true) { "`n$($SubMsg6)" | Add-Content $logFile }
Write-Host $SubMsg6

#Get the CSV file
$CSVFile2 = Import-Csv $CSVFileSubs

#Read CSV file and create document library indexes
ForEach($Line2 in $CSVFile2) {

    Try {

    ## Build URL
    $SubsiteUrl2 = $Line2.SubsiteUrl
    $SiteName2 = $Line2.ParentUrl
    $SiteURL2 = "https://$tenantName.sharepoint.com$SiteName2"
    $FullSiteURL2 = "https://$tenantName.sharepoint.com$SiteName2" + "/" + $SubsiteUrl2

    $SubMsg7 = "   --  Indexing Document Library on: " + $FullSiteURL2 + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg7)" | Add-Content $logFile }
    Write-Host $SubMsg7
    ## Set Library Name
    #$LibraryName = $Line.Library
    $LibraryName = "Documents"

    #Connect to PNP Online
    Connect-PnPOnline -Url $FullSiteURL2 -Credential $TenantCredential
    Write-Host " "
    Write-Host "PnP Service: CONNECTED" -ForegroundColor Green
 
    #Get the Context
    $Context = Get-PnPContext

    #Get the Field from List
    $Field1 = Get-PnPField -List $LibraryName -Identity $ColumnName1
    $Field2 = Get-PnPField -List $LibraryName -Identity $ColumnName2
    $Field3 = Get-PnPField -List $LibraryName -Identity $ColumnName3
    $Field4 = Get-PnPField -List $LibraryName -Identity $ColumnName4
    $Field5 = Get-PnPField -List $LibraryName -Identity $ColumnName5
    $Field6 = Get-PnPField -List $LibraryName -Identity $ColumnName6
 
    #Set the Indexed Property of the Field
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName1
    $Field1.Indexed = $True
    $Field1.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName2
    $Field2.Indexed = $True
    $Field2.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName3
    $Field3.Indexed = $True
    $Field3.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName4
    $Field4.Indexed = $True
    $Field4.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName5
    $Field5.Indexed = $True
    $Field5.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName6
    $Field6.Indexed = $True
    $Field6.Update() 

    $Context.ExecuteQuery()
    Write-Host " "
    $SubMsg8 = "   --  Columns Indexed: $($ColumnName1),  $($ColumnName2), $($ColumnName3), $($ColumnName4), $($ColumnName5), $($ColumnName6) - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg8)" | Add-Content $logFile }
    $SubMsg9 = "   --  Column Indexing Complete for 'Documents' Library on " + $FullSiteURL2 + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg9)" | Add-Content $logFile }
    Write-Host $SubMsg9 -ForegroundColor Green

    ## Pause Briefly
    Start-Sleep -Seconds 5

    }
    
    Catch {
        $SubMsg10 = "   ##  Error Indexing Column! " + $_.Exception.Message + " - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($SubMsg10)" | Add-Content $logFile }
        write-host $SubMsg10 -f Red
          }

}


    ## Disconnect from PnP Connection
    Disconnect-PnPOnline

    ## Log Complete
    $SubMsg11 = "*** Subsite Build and Index Operation Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($SubMsg11)" | Add-Content $logFile }
    Write-Host $SubMsg11 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B4"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C4"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate3 = "Subsite Build and Index has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate3)" | Add-Content $logFile }

    pause
    CallMenu

}

## 4) Library Build & Index
function CmdLibs {
    $CSVFileLibs = "$TempDir\Libraries.csv"

    #Get the CSV file
    $CSVFile3 = Import-Csv $CSVFileLibs

    ## Begin Subsite Build
    $LibMsg1 = "-- BEGIN Document Library Builds - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg1)" | Add-Content $logFile }

#Read CSV file and create document library indexes
ForEach($Line3 in $CSVFile3) {

    Try {

    ## Build URL
    $SiteName = $Line.SiteLocation
    ## ROOT SITE
    #$SiteURL = "https://$tenantName.sharepoint.com$SiteName"
    ## Alternate Site Collection
    $SiteURL = "https://$tenantName.sharepoint.com/sites$SiteName"
    $LibMsg2 = "  -- Building SP Libraries at site: " + $SiteURL + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg2)" | Add-Content $logFile }
    Write-Host $LibMsg2
    ## Set Library Variables
    $LibraryName = $Line3.Library
    $LibraryUrl = $Line3.SiteLocation
    $LibraryTemplate = $Line3.Template

    #Connect to PNP Online
    Connect-PnPOnline -Url $SiteURL -Credential $TenantCredential
    Write-Host " "
    Write-Host "PnP Service: CONNECTED" -f Green

    ## Library Being Created
    Write-Host " "
    $LibMsg3 = "   -- Library being created and Indexed: " + $LibraryName + "- $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg3)" | Add-Content $logFile }
    $LibMsg4 = "    -- Details: " + $LibraryName + " | " + $LibraryUrl + " | " + $LibraryTemplate + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg4)" | Add-Content $logFile }
    Write-Host $LibMsg3 -f Green
    Write-Host $LibMsg4 -f Yellow
 
    ## FIRST: Create the destination Library
    New-PnPList -Title $Line3.Library -Url $Line3.SiteLocation -Template $Line3.Template -OnQuickLaunch

    ## Pause briefly while Library is created
    Write-Host " "
    Write-Host "Pause While Library is created..." -f Yellow
    Start-Sleep -Seconds 5

    ## SECOND: Create the destination Library
    ## Begin Indexing
    Write-Host " "
    Write-Host "Begin Library Index Settings"

    #Get the Context
    $Context = Get-PnPContext
    
    #Get the Field from List
    $Field1 = Get-PnPField -List $LibraryName -Identity $ColumnName1
    $Field2 = Get-PnPField -List $LibraryName -Identity $ColumnName2
    $Field3 = Get-PnPField -List $LibraryName -Identity $ColumnName3
    $Field4 = Get-PnPField -List $LibraryName -Identity $ColumnName4
    $Field5 = Get-PnPField -List $LibraryName -Identity $ColumnName5
    $Field6 = Get-PnPField -List $LibraryName -Identity $ColumnName6
 
    #Set the Indexed Property of the Field
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName1
    $Field1.Indexed = $True
    $Field1.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName2
    $Field2.Indexed = $True
    $Field2.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName3
    $Field3.Indexed = $True
    $Field3.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName4
    $Field4.Indexed = $True
    $Field4.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName5
    $Field5.Indexed = $True
    $Field5.Update() 
    Write-Host " "
    Write-Host "--- Setting index for Column: " $ColumnName6
    $Field6.Indexed = $True
    $Field6.Update() 

    $Context.ExecuteQuery()
    Write-Host " "
    $LibMsg5 = "   -- Columns Indexed: $($ColumnName1),  $($ColumnName2), $($ColumnName3), $($ColumnName4), $($ColumnName5), $($ColumnName6) - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg5)" | Add-Content $logFile }
    $LibMsg6 = "   -- Column Indexing Complete for 'Documents' Library on " + $FullSiteURL2 + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg6)" | Add-Content $logFile }
    Write-Host $LibMsg6 -ForegroundColor Green

    Start-Sleep -Seconds 5

    ## Disconnect from PnP Connection
    Disconnect-PnPOnline

    }
    
    Catch {
        $LibMsg7 = "   ##  Error Indexing Column! " + $_.Exception.Message + " - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($LibMsg7)" | Add-Content $logFile }
        write-host $LibMsg7 -f Red
          }

}

    ## Log Complete
    $LibMsg8 = "*** Library Builds and Indexing Operation Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($LibMsg8)" | Add-Content $logFile }
    Write-Host $LibMsg8 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B5"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C5"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate4 = "Libraries Build and Index has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate4)" | Add-Content $logFile }

    pause
    CallMenu
    
}

## 5) Preprovision OneDrive Sites
function CmdODfB {
    ## User List
    $users = Get-Content -path "$TempDir\TenantUsers.txt"


    ## Select SINGLE User [S] or LIST of Users [L]
    Write-Host " "
    $SiteInput = Read-Host "Is a LIST of users [L] or a SINGLE user [S] to Provision?"
    if ($SiteInput -eq 'L')
    {
    $UserList = $users
    $Listmsg1 = "You've selected to pre-provision OneDrive for Business users from a LIST. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Listmsg1)" | Add-Content $logFile }
    Write-Host " "
    Write-Host $Listmsg1 -ForegroundColor Yellow

    }
    elseif ($SiteInput -eq 'S')
    {
    Write-Host " "
    $Singlemsg1 = "You've selected to pre-provision OneDrive for Business site for a SINGLE user. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Singlemsg1)" | Add-Content $logFile }
    Write-Host $Singlemsg1 -ForegroundColor Yellow
    Write-Host "Enter the UPN (Email or .onmicrosoft.com address) for this single user." -ForegroundColor Green
    $UserList = Read-Host "--- Enter UPN for user ---> "
    $Singlemsg2 = "The SINGLE user is: " + $UserList + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Singlemsg2)" | Add-Content $logFile }
    Write-Host $Singlemsg2

    }else
    {
    Write-Host " "
    Write-Output "Please only type in 'R' or 'S'."
    }


    ## Begin Preprovision Process
    $ODfBMsg1 = "BEGIN Pre-Provision Process - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ODfBMsg1)" | Add-Content $logFile }


    ## Connecting Dialog
    Write-Host "Connecting to SPO Admin Site "$(Get-TimeStamp) -ForegroundColor Yellow
    Write-Host "   ...  ...  ...  ...  ...  ...  ...  ...  ...  ..."
    Write-Host " "

    ## Connect to SPO Admin Site
    Connect-SPOService -Url $tenantAdminUrl -Credential $TenantCredential

    ## Connected
    Write-Host "Connected..."$(Get-TimeStamp) -ForegroundColor Yellow
    Write-Host "   ...  ...  ...  ...  ...  ...  ...  ...  ...  ..."
    Write-Host " "

    ## Document User Input list OR single user
    $ODfBMsg2 = "The USER(S) is documented below. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ODfBMsg2)" | Add-Content $logFile }
    $ODfBMsg3 = $UserList
    if ($log -eq $true) { "`n$($ODfBMsg3)" | Add-Content $logFile }

    $ODfBMsg4 = "Requesting Personal Site - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ODfBMsg4)" | Add-Content $logFile }

    Try {
    ## Process each user in the list and request One Drive site
    Request-SPOPersonalSite -UserEmails $users -NoWait
    }
    Catch {
            $ODfBMsg5 = " ##  Error Provisioning ODfB site(s)! " + $_.Exception.Message + " - $(Get-TimeStamp)"
            if ($log -eq $true) { "`n$($ODfBMsg5)" | Add-Content $logFile }
            write-host $ODfBMsg5 -f Red
    }

    ## One Drives Sites Requested
    $ODfBMsg6 = "One Drive sites requested successfully. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ODfBMsg6)" | Add-Content $logFile }
    Write-Host $ODfBMsg6 -ForegroundColor Yellow
    Write-Host "   ...  ...  ...  ...  ...  ...  ...  ...  ...  ..."
    Write-Host " "

    ## Disconnect from SPO Service
    Disconnect-SPOService

    ## Let user know script has completed
    Write-Host "   *** " -ForegroundColor Yellow -NoNewline
    Write-Host "One Drive Sites Successfully Requested." -ForegroundColor Green -NoNewline
    Write-Host " ***" -ForegroundColor Yellow -NoNewLine
    Write-Host "   "$(Get-TimeStamp)

    ## Log Complete
    $ODfBMsg7 = "*** OneDrive for Business Pre-Provision Operations Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ODfBMsg7)" | Add-Content $logFile }
    Write-Host $ODfBMsg7 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B6"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C6"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate5 = "OneDrive for Business Pre-Provision has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate5)" | Add-Content $logFile }

    pause
    CallMenu

}

## 6) Allow Scripts on Site
function CmdAllowScripts {

    ## Select Root Site or other Site Collection to use
    Write-Host "Which site are you setting to allow scripts to run? "
    $SiteInput = Read-Host "Is This Root Site [R] or another Site Collection [S]?"
    if ($SiteInput -eq 'R')
    {
    $SiteURL = "https://$tenantName.sharepoint.com"
    $Rootmsg1 = "You've selected to ALLOW SITE SCRIPTING at the root site collection. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Rootmsg1)" | Add-Content $logFile }
    $Rootmsg2 = "The Root Site URL is: " + $SiteURL + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Rootmsg2)" | Add-Content $logFile }
    Write-Host " "
    Write-Host $Rootmsg1 -f Yellow
    Write-Host $Rootmsg2

    }
    elseif ($SiteInput -eq 'S')
    {
    Write-Host " "
    $Sitemsg1 = "You've selected to ALLOW SITE SCRIPTING at a site OTHER than the root site collection. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Sitemsg1)" | Add-Content $logFile }
    Write-Host $Sitemsg1 -ForegroundColor Yellow
    Write-Host "Enter the Site URL, only the part AFTER .../sites/ In the URL" -ForegroundColor Green
    $SiteName = Read-Host "--- Enter name of the Site Collection ---> "
    $SiteURL = "https://$tenantName.sharepoint.com/sites/$SiteName"
    $Sitemsg2 = "The Site URL is: " + $SiteURL + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($Sitemsg2)" | Add-Content $logFile }
    Write-Host "The Site URL is: " $SiteURL

    }else
    {
    Write-Host " "
    Write-Output "Please only type in 'R' or 'S'."
    }


    ##Variables for Admin Center & Site Collection URL
    $tenantAdminUrl = "https://$tenantName-admin.sharepoint.com/"

    ## Connect to SharePoint Online
    Connect-SPOService -Url $tenantAdminUrl -Credential $TenantCredential

    ## Connect Log
    $AllowScriptMsg1 = "Connected to " + $tenantAdminUrl + "with user " + $tenantAdmin + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($AllowScriptMsg1)" | Add-Content $logFile }
 
    ## Allow/Deny DenyAddAndCustomizePages Flag -- Allow = $False / Deny = $True
    Try {
        Set-SPOSite $SiteURL -DenyAddAndCustomizePages $False
        $AllowScriptMsg2 = " -- Scripting set to ALLOW for site: " + $SiteURL + " - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($AllowScriptMsg2)" | Add-Content $logFile }
        Write-Host " "
        write-host $AllowScriptMsg2 -f Yellow 
        }

    Catch {
        $AllowScriptMsg3 = " ## Error allowing scripting for: " + $SiteURL + " -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($AllowScriptMsg3)" | Add-Content $logFile }
        Write-Host " "
        write-host $AllowScriptMsg3 -f Red 
          }


    ## Disconnect SPO Service
    Disconnect-SPOService

    $AllowScriptMsg4 = "Disconnected from " + $tenantAdminUrl + " - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($AllowScriptMsg4)" | Add-Content $logFile }

    ## Log Complete
    $AllowScriptMsg5 = "*** 'Allow Scripting' Operations Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($AllowScriptMsg5)" | Add-Content $logFile }
    Write-Host $AllowScriptMsg5 -f Green


    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B7"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C7"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate6 = "Create Navigation has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate6)" | Add-Content $logFile }

    pause
    CallMenu


}

## 7) LISTS - Save as Templates
function CmdListTemp {

    ## Connect Log
    $ListSave1 = "Ready to Save Lists to the List Template Gallery. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ListSav1)" | Add-Content $logFile }
    Write-Host $ListSave1

    
    $CSVFileLists = "$TempDir\Lists.csv"

    #Get the CSV file
    $CSVFile4 = Import-Csv $CSVFileLists


    #Read CSV file and save all lists as tempates
    ForEach($Line4 in $CSVFile4) {
 
    Try{
   
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Line4.SourceSiteUrl)
        $Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($TenantAdmin, $TenantAdminPw)
  
        #Get the List
        $List = $Ctx.Web.lists.GetByTitle($Line4.ListName)
        $List.SaveAsTemplate($Line4.TemplateFileName, $Line4.TemplateName, $Line4.TemplateDescription, $IncludeData)
        $Ctx.ExecuteQuery()
 
        #Write-Host -f Green "List $Line.ListName Saved as Template!"
        ## Debug
        $ListSave2 = " -- List [ID: " + $Line4.ID + " | Name: " + $Line4.ListName + " ] SAVED as template! - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($ListSave2)" | Add-Content $logFile }
        Write-Host $ListSave2 -f Green


    }
    Catch {
            $ListSave3 = " ## Error Saving List [ID: " + $Line4.ID + " | Name: " + $Line4.ListName + " ] as template! -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
            if ($log -eq $true) { "`n$($ListSave3)" | Add-Content $logFile }
            Write-Host " "
            write-host $ListSave3 -f Red 
        }

    }

    ## Log Complete
    $ListSave4 = "*** List Template Save Operations Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ListSave4)" | Add-Content $logFile }
    Write-Host $ListSave4 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B8"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C8"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate7 = "Save Lists as Templates has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate7)" | Add-Content $logFile }

    pause
    CallMenu


}

## 8) LISTS - Download Templates
function CmdDownTemp {

    ## Connect Log
    $TemplateDL1 = "Ready to Download List Templates from List Template Gallery. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($TemplateDL1)" | Add-Content $logFile }
    Write-Host $TemplateDL1

    $CSVFileTemplates = "$TempDir\Lists.csv"

    #Get the CSV file
    $CSVFile5 = Import-Csv $CSVFileTemplates

    ForEach($Line5 in $CSVFile5)
{
 
    Try{
   
        ## Variables
        $ExportFile = "$TempDir\" + $Line.TemplateFileName
        $SiteURL = $Line5.SourceSiteUrl
        $ListTemplateName = $Line5.TemplateName

        ## Call the function to Download the list template
        Download-SPOListTemplate -SiteURL $SiteURL -ListTemplateName $ListTemplateName -ExportFile $ExportFile
        $TemplateDL2 = " -- Downloaded List Template [ID: " + $Line5.ID + "Name: " + $Line5.TemplateFileName + " ] successfully! - $(Get-TimeStamp)"
        if ($log -eq $true) { "`n$($TemplateDL2)" | Add-Content $logFile }
        Write-Host $TemplateDL2 -f Green

    }
    Catch {
            $TemplateDL3 = " ## Error Downloading Template File [ID: " + $Line5.ID + " Name: " + $Line5.ListName + " ] as template! -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
            if ($log -eq $true) { "`n$($TemplateDL3)" | Add-Content $logFile }
            write-host $TemplateDL3 -f Red
    }

    }

    ## Log Complete
    $TemplateDL4 = "*** List Template Download Operations Complete. *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($TemplateDL4)" | Add-Content $logFile }
    Write-Host $TemplateDL4 -f Green


    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B9"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C9"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate8 = "Download and Save List Templates has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate8)" | Add-Content $logFile }

    pause
    CallMenu


}

## 9) LISTS - Upload and Create
function CmdListCreate {

    $ListBuild1 = "Begin IMPORTING List Template Files to Template Gallery... - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ListBuild1)" | Add-Content $logFile }
    Write-Host $ListBuild1 -f Yellow
    

    ###################################################################
    #                       UPLOAD TEMPLATES                          #
    ###################################################################

    $CSVListBuild = "$TempDir\Lists.csv"

    #Get the CSV file
    $CSVFile6 = Import-Csv $CSVListBuild

    #Read CSV file and save all lists as tempates
    ForEach($TemplateLine in $CSVFile6)
    {
 
    Try {
        
            $DestSiteColUrl = $TemplateLine.DestSiteUrl
            #Get Credentials to connect
            $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($TenantAdmin, $TenantAdminPw)
  
            #Setup the context
            $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($DestSiteColUrl)
            $Ctx.Credentials = $Credentials

            #Set Import Variables
            $TemplateName = $TemplateLine.TemplateName
            $ImportTemplateName = $TemplateLine.TemplateName
            $ImportFileName = $TemplateLine.TemplateFileName
            $ImportFile = "$TempDir\" + $ImportFileName
         
            #Get the "List Templates" Library
            $List= $Ctx.web.Lists.GetByTitle("List Template Gallery")
            $ListTemplates = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
            $Ctx.Load($ListTemplates)
            $Ctx.ExecuteQuery()
 
            #Check if the Given List Template already exists
            $ListTemplate = $ListTemplates | where { $_["TemplateTitle"] -eq $TemplateName }
 
            If($ListTemplate -eq $Null)
            {
                #Get the file from disk
                $FileStream = ([System.IO.FileInfo] (Get-Item $ImportFile)).OpenRead()
                #Get File Name from source file path
                $TemplateFileName = Split-path $ImportFile -leaf
    
                #Upload the File to SharePoint Library
                $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                $FileCreationInfo.Overwrite = $true
                $FileCreationInfo.ContentStream = $FileStream
                $FileCreationInfo.URL = $TemplateFileName
                $FileUploaded = $List.RootFolder.Files.Add($FileCreationInfo)
                $Ctx.Load($FileUploaded)
                $Ctx.ExecuteQuery()
             
                #Set Metadata of the File
                $ListItem = $FileUploaded.ListItemAllFields
                $Listitem["TemplateTitle"] = $TemplateName
                $Listitem["FileLeafRef"] = $ImportFileName
                $ListItem.Update()
                $Ctx.ExecuteQuery()
  
                #Close file stream
                $FileStream.Close()
 
                $ListBuild2 = " -- List Template '$ImportFileName' Uploaded - $(Get-TimeStamp)"
                if ($log -eq $true) { "`n$($ListBuild2)" | Add-Content $logFile }
                Write-Host $ListBuild2 -f Green

            }
            else
            {
                $ListBuild3 = " -- List Template '$ImportFileName' Already Exists - $(Get-TimeStamp)"
                if ($log -eq $true) { "`n$($ListBuild3)" | Add-Content $logFile }
                Write-Host $ListBuild3 -f Green
            }
        }
        Catch {
            $ListBuild4 = " ## Error Uploading List Template! -- " + $_.Exception.Message + " - $(Get-TimeStamp)"
            if ($log -eq $true) { "`n$($ListBuild4)" | Add-Content $logFile }
            write-host $ListBuild4 -f Red
        }
    Write-Host -f White ". . ."
    }

    $ListBuild5 = " -- List Template Upload operations completed. - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ListBuild5)" | Add-Content $logFile }
    Write-Host $ListBuild5 -f Green

    Write-Host -f White "Brief Pause Before Creating Lists from Template."
    Write-Host -f White ". . . . ."
    Sleep -Seconds 7
    Write-Host " "

    $ListBuild6 = " -- Begin creating Lists from Templates... - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ListBuild6)" | Add-Content $logFile }
    Write-Host $ListBuild6 -f Yellow
    Write-Host " "

    Sleep -Seconds 3

    ###################################################################
    #                       CREATE LISTS FROM TEMPLATE GALLERY        #
    ###################################################################

    #Read CSV file and save all lists as tempates
    ForEach($Line6 in $CSVFile6)
    {

        #Setup Working Variables
        $SiteUrl = $Line6.DestSiteUrl
        $ListTemplateName = $Line6.TemplateName
        $ListName = $Line6.ListName


        #Get Credentials to connect
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($TenantAdmin, $TenantAdminPw)
   
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
 
        #Get All Lists
        $Lists = $Ctx.Web.Lists
        $Ctx.Load($Lists)
        $Ctx.ExecuteQuery()
 
        #Get the Custom list template
        $ListTemplates=$Ctx.site.GetCustomListTemplates($Ctx.site.RootWeb)
        $Ctx.Load($ListTemplates)
        $Ctx.ExecuteQuery()
 
        #Filter Specific List Template
        $ListTemplate = $ListTemplates | where { $_.Name -eq $ListTemplateName }
        If($ListTemplate -ne $Null)
        {
            #Check if the given List exists
            $List = $Lists | where {$_.Title -eq $ListName}
            If($List -eq $Null)
            {
                #Create new list from custom list template
                $ListCreation = New-Object Microsoft.SharePoint.Client.ListCreationInformation
                $ListCreation.Title = $ListName
                $ListCreation.ListTemplate = $ListTemplate
                $List = $Lists.Add($ListCreation)
                $Ctx.ExecuteQuery()

                $ListBuild7 = " -- List $ListName Created from Custom List Template Successfully! - $(Get-TimeStamp)"
                if ($log -eq $true) { "`n$($ListBuild7)" | Add-Content $logFile }
                Write-Host $ListBuild7 -f Green

            }
        else
            {
                $ListBuild8 = " -- List '$($ListName)' Already Exists! - $(Get-TimeStamp)"
                if ($log -eq $true) { "`n$($ListBuild8)" | Add-Content $logFile }
                Write-Host $ListBuild8 -f Yellow
            }
            }
            else
                {
                $ListBuild9 = " -- List Template '$($ListTemplateName)' Not Found! - $(Get-TimeStamp)"
                if ($log -eq $true) { "`n$($ListBuild9)" | Add-Content $logFile }
                Write-Host $ListBuild9 -f Yellow
                }


        Write-Host -f White ". . ."
    }

    Write-Host " "
    ## Log Complete
    $ListBuild10 = "*** List Creation Operations Complete! *** - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ListBuild10)" | Add-Content $logFile }
    Write-Host $ListBuild10 -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B10"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C10"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate9 = "Upload Templates and Create Lists has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate9)" | Add-Content $logFile }

    pause
    CallMenu


}

## 10) Assign Permissions
function CmdPerms {

    Write-Host "Place holder for '10) Assign Permissions'." -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B11"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C11"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate10 = "Assign Permissions has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate10)" | Add-Content $logFile }

    pause
    CallMenu


}

## 11) Export Navigation
function CmdExportNav {
  
   $ExportNav1 = "Begin EXPORTING Site Navigation elements... - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExportNav1)" | Add-Content $logFile }
    Write-Host $ExportNav1 -f Yellow

    ###################################################################
    #            GATHER SOURCE NAV SITE CREDENTIALS                   #
    ###################################################################

    ## Gather Creds
    $SrcNavTenantAdmin = Read-Host "--- Enter Tenant Admin from Source Nav Site --->  "
    $SrcNavTenantAdminPw = Read-Host "--- Enter Tenant Admin Password from Source Nav Site --->" -AsSecureString

    ## Gather Tenant Name
    $SrcNavUrl = Read-Host "--- Enter URL to EXPORT site navigation ---> "

    $ExportNav2 = " -- You chose to EXPORT site navigation elements from site: $SrcNavUrl  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExportNav2)" | Add-Content $logFile }
    Write-Host $ExportNav2 -f Green

    Write-Host " "
    Write-Host "Which navigation type will you be exporting?   " -f Yellow -NoNewline
    Write-Host "[Default=TopNavigationBar]"
    Write-Host "  TopNavigationBar, Footer, QuickLaunch, SearchNav"
    $SrcNavLocation = Read-Host "--- Navigation Source Type --->  "
    if ([string]::IsNullOrWhiteSpace($SrcNavLocation))

    {

    $SrcNavLocation = "TopNavigationBar"

    }

    Write-Host " "
    $ExportNav3 = " -- You chose to EXPORT site navigation type: $SrcNavLocation  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExportNav3)" | Add-Content $logFile }
    Write-Host $ExportNav3 -f Green
   
   ## Call Export
   ExportSiteNavigation -SourceSite $SrcNavUrl -NavigationLocation $SrcNavLocation -BackupDestination $NavExportFile


    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B12"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C12"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate11 = "EXPORT Navigation has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate11)" | Add-Content $logFile }

    pause
    CallMenu
   
     
}
    
## 12) Import Navigation
function CmdImportNav {

    $ImportNav1 = "Begin IMPORTING Site Navigation elements... - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav1)" | Add-Content $logFile }
    Write-Host $ImportNav1 -f Yellow
    
    $NavExportFile = "$TempDir\SiteNavigationBackup.xlsx"
    #$NavExportFile = ".\SiteNavigationBackup.xlsx"
 
    ## Gather Tenant Name
    $DestNavUrl = Read-Host "--- Enter URL to IMPORT site navigation into ---> "

    $ImportNav2 = " -- You chose to IMPORT site navigation elements from site: $DestNavUrl  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav2)" | Add-Content $logFile }
    Write-Host $ImportNav2 -f Green

    Write-Host " "
    Write-Host "Which navigation type will you be importing?   " -f Yellow -NoNewline
    Write-Host "[Default=TopNavigationBar]"
    Write-Host "  TopNavigationBar, Footer, QuickLaunch, SearchNav"
    $DestNavLocation = Read-Host "--- Navigation Source Type --->  "
    if ([string]::IsNullOrWhiteSpace($DestNavLocation))

    {

    $DestNavLocation = "TopNavigationBar"

    }

    Write-Host " "
    $ImportNav3 = " -- You chose to IMPORT site navigation type: $DestNavLocation  - $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ImportNav3)" | Add-Content $logFile }
    Write-Host $ImportNav3 -f Green

    ## Call Import
    ImportSiteNavigation -DestinationSite $DestNavUrl -NavigationLocation $DestNavLocation

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B13"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C13"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate12 = "IMPORT Navigation has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate12)" | Add-Content $logFile }

    pause
    CallMenu


}

## 13) UTIL:Add Admin to ALL Sites
function CmdAddAdmins {

    Write-Host "Place holder for '13) UTIL:Add Admin to ALL Sites'." -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B14"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C14"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate13 = "Assign Permissions has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate13)" | Add-Content $logFile }

    pause
    CallMenu

}

## 14) UTIL:Remove Admin from ALL Sites
function CmdRemoveAdmins {

    Write-Host "Place holder for '14) UTIL:Remove Admin from ALL Sites'." -f Green

    ## Checklist Update
    $UpdateComplete = 'Yes'
    $ExcelPath = $ExportFile
    $excel = Open-ExcelPackage -Path $ExcelPath
    $excel.'Migration Checklist'.Cells["B15"].Value = $UpdateComplete
    $excel.'Migration Checklist'.Cells["C15"].Value = Get-TimeStamp
    Close-ExcelPackage -ExcelPackage $excel
    $ExcelUpdate14 = "Assign Permissions has been updated in the checklist to Completed = 'Yes' with timestamp: $(Get-TimeStamp)"
    if ($log -eq $true) { "`n$($ExcelUpdate14)" | Add-Content $logFile }

    pause
    CallMenu

}

## FUNCTION: Build Checklist Worksheet
function BuildChecklist {

$ExcelParams = @{
    #Path      = $env:TEMP + '\Excel.xlsx'
    #Path      = 'C:\Temp\Excel-Test6.xlsx'
    Path       = $ExportFile
    #Show      = $false
    Verbose   = $false
}
Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
$Array = @()
$Step01 = [PSCustomObject]@{
    MigrationStep   = 'Allow PnP Auth'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'The script allows PnP cmdlets to be run against the tenant.'
}

$Step02 = [PSCustomObject]@{
    MigrationStep   = 'Site Collection Build'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Takes CSV file input to use as build manifest for all site collections.'
}
  
$Step03 = [PSCustomObject]@{
    MigrationStep   = 'Sub Site Build'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Build subsites on the previously built site collections.'
}
  
$Step04 = [PSCustomObject]@{
    MigrationStep   = 'Library Build/Index'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Takes CSV file input as build manifest to build and then index libraries.'
}

$Step05 = [PSCustomObject]@{
    MigrationStep   = 'Pre-Provision OneDrive'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Uses TXT file of user accounts/UPNs to pre-provision destination OneDrive sites.'
}

$Step06 = [PSCustomObject]@{
    MigrationStep   = 'Allow Scripting on Site'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Checks the option for "Allow Custom Scripting" on a site collection. Used to expose List Template Gallerys and the "Save as Template" option.'
}

$Step07 = [PSCustomObject]@{
    MigrationStep   = 'LISTS - Save as Templates'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Uses CSV input to save lists as templates in the List Template Gallery on each site collection. Requires "Allow Scripting" script run.'
}

$Step08 = [PSCustomObject]@{
    MigrationStep   = 'LISTS - Download Templates'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Downloads List templates to local directory from the List Template Galleryon a site.'
}

$Step09 = [PSCustomObject]@{
    MigrationStep   = 'LISTS - Upload and Create'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Uploads .stp template files to List Template Gallery and build List based on selected template.'
}

$Step10 = [PSCustomObject]@{
    MigrationStep   = 'Assign Permissions'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Assigns permissions to destination sites/libraries.'
}

$Step11 = [PSCustomObject]@{
    MigrationStep   = 'EXPORT Navigation'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Exports custom navigation for source site.'
}

$Step12 = [PSCustomObject]@{
    MigrationStep   = 'IMPORT Navigation'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Imports custom navigation for destination site.'
}
$Step13 = [PSCustomObject]@{
    MigrationStep   = 'Add Admin to ALL Site Collection'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Utility function:  ADDS current admin account as site owner to all site collections.'
}
$Step14 = [PSCustomObject]@{
    MigrationStep   = 'Add Admin to ALL Site Collection'
    Completed   = 'NO'
    Time = 'NOT STARTED'
    Description = 'Utility function:  REMOVES current admin account as site owner from all site collections.'
}


$Array = $Step01, $Step02, $Step03, $Step04, $Step05, $Step06, $Step07, $Step08, $Step09, $Step10, $Step11, $Step12, $Step13, $Step14
#$Array | Out-GridView -Title 'Migration Checklist'
#$Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -WorksheetName Numbers
$Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -AutoSize -TableName "Checklist" -WorksheetName "Migration Checklist" -Append -ConditionalText $(
                New-ConditionalText 'NO' DarkRed LightPink
                New-ConditionalText 'NOT STARTED' DarkRed LightPink
                New-ConditionalText 'Yes' Black LightGreen
                New-ConditionalText 'N/A' Black LightYellow
                )

}

## Delay and Wait Function -- 15 Second Delay
function DelayAndWait
{
Write-Host " "
Write-Host "Please wait while the process completes..."
Start-Sleep -Seconds 3
Write-Host " ." -f Yellow -NoNewline
Start-Sleep -Seconds 3
Write-Host " . ." -f Yellow -NoNewline
Start-Sleep -Seconds 3
Write-Host " . . ." -f Yellow -NoNewline
Start-Sleep -Seconds 3
Write-Host " . . . ." -f Yellow -NoNewline
Start-Sleep -Seconds 3
Write-Host " . . . . ." -f Yellow

#Read-Host "Press ENTER to continue..."
}

## Log ON ($true) or OFF ($false)
$log = $true

#pause

###################################################################
#            BUILD TEXT INTERFACE -- GATHER CREDS                 #
###################################################################

## Clear Screen
CLS

## Call Interface Function
MainHeaderLayout

###################################################################
#            GATHER CREDENTIALS                                   #
###################################################################

## Gather Tenant Name
$tenantName = Read-Host "--- Enter name of Tenant ---> "

## Gather Creds, must be licensed for Power BI
$TenantAdmin = Read-Host "--- Enter Tenant Admin --->  "
$TenantAdminPw = Read-Host "--- Enter Tenant Admin Password --->" -AsSecureString

## Build Credential to Connect
$TenantCredential = New-Object System.Management.Automation.PSCredential($TenantAdmin,$TenantAdminPw)

## Get Credentials to Connect
#$Credential = Get-Credential

## Set Tenant Admin URL
$tenantAdminUrl = "https://$tenantName-admin.sharepoint.com"


###################################################################
#            SECONDARY FUNCTIONS                                  #
###################################################################

## Create Temp directory "C:\Temp-TenantName"
$TempDir = "C:\Temp-$tenantName"

## Function to Check if temp directory exists, and create if it does not
if (-not (Test-Path -LiteralPath $TempDir)) {
    
    try {
        New-Item -Path $TempDir -ItemType Directory -ErrorAction Stop | Out-Null #-Force
    }
    catch {
        Write-Error -Message "Unable to create directory '$TempDir'. Error was: $_" -ErrorAction Stop
    }
    "Successfully created directory '$TempDir'."

}
else {
    "Directory already exists."
}


## Set Current Working Directory
Set-Location -Path $TempDir -PassThru

## Set file for export results
$ExportFile = ".\MigrationCheckList_" + $tenantName + ".xlsx"

## See if checklist spreadsheet exists, if not create if it does not
if (-not (Test-Path -LiteralPath $ExportFile)) {
    
    try {
        ## BuildChecklist Spreadsheet in Temp directory
        BuildChecklist
    }
    catch {
        Write-Error -Message "Unable to create Migration Checklist '$ExportFile'. Error was: $_" -ErrorAction Stop
    }
    "Successfully created Migration Checklist '$ExportFile'."

}
else {
    "Migration Checklist already exists."
}

###################################################################
#            LOGGING OPERATIONS                                   #
###################################################################

if ($log -eq $true)

{

  ## Construct a log file name based on the date that
  ## we can save progress to
  $logStart = Get-Date
  $logStartDate = "$($logStart.Year)-$($logStart.Month)-$($logStart.Day)"
  $logStartTime = "$($logStart.Hour)-$($logStart.Minute)-$($logStart.Second)"
  $logFile = ".\$tenantName-Migration_" + $logStartDate + "-" + $logStartTime + ".txt"

}

## Log Header
$varmsg = "OPERATION: SPO Admin Tool for: $tenantName | Script Revised: " + $RevDate
if ($log -eq $true) { "`n$($varmsg)" | Add-Content $logFile }

###################################################################
#            PREREQUISITE CHECK                                   #
###################################################################

## Prereq Check
CheckModPnP

###################################################################
#            LOAD CSV'S -- PROMPTS                                #
###################################################################

## Open Explorer to $Tempdir
explorer.exe $TempDir

## Prompt to load CSV's to TempDir
Write-Host " "
Write-Host "*** " -f Green -NoNewline
Write-Host "Before continuing, please copy the CSV files for the build operations to: $TempDir then hit 'ENTER'." -f Yellow -NoNewline
Write-Host " ***" -f Green
Write-Host " "
Write-Host "Required file(s):  " -NoNewline
Write-Host "'SiteCollections.csv', 'Subsites.csv', 'Libraries.csv', 'TenantUsers.txt', 'Lists.csv' " -f Yellow
$FileLoad1 = "Prompt for CSV load to $Tempdir. - $(Get-TimeStamp)"
if ($log -eq $true) { "`n$($FileLoad1)" | Add-Content $logFile }
Pause


###################################################################
#            BUILD TEXT INTERFACE -- BUILD FUNCTIONS              #
###################################################################

## BuildChecklist Spreadsheet in Temp directory
#BuildChecklist

## Clear Screen
CLS

## Call Interface Function
MainHeaderLayout

## Call Menu
CallMenu

## Main Menu Options Loop & Actions
Do{
    
    $MenuResponse = Read-Host "Choose Operation"

Switch($MenuResponse){
    '1' {
       #1) PnP Auth
       $MenuSelect1 = "** Running option 1: PnP Auth  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect1)" | Add-Content $logFile }
       Write-Host $MenuSelect1 -f Yellow
       CmdPnPAuth
       CallMenu
      }
    '2' {
       #2) Site Collection Build
       $MenuSelect2 = "** Running option 2: Site Collection Build  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect2)" | Add-Content $logFile }
       Write-Host $MenuSelect2 -f Yellow
       CmdscBuild  
       CallMenu

      }
    '3' {
       #3) Subsite Build
       $MenuSelect3 = "** Running option 3: Subsite Build  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect3)" | Add-Content $logFile }
       Write-Host $MenuSelect3 -f Yellow 
       CmdSubBuild
       CallMenu
        
      }
    '4' {
       #4) Library Build & Index
       $MenuSelect4 = "** Running option 4: Library Build & Index  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect4)" | Add-Content $logFile }
       Write-Host $MenuSelect4 -f Yellow 
       CmdLibs
       CallMenu

      }
    '5' {
       #5) Preprovision OneDrive Sites
       $MenuSelect5 = "** Running option 5: Preprovision OneDrive Sites  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect5)" | Add-Content $logFile }
       Write-Host $MenuSelect5 -f Yellow
       CmdODfB 
       CallMenu

      }
    '6' {
       #6) Allow Scripts on Site
       $MenuSelect6 = "** Running option 6: Allow Scripts on Site  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect6)" | Add-Content $logFile }
       Write-Host $MenuSelect6 -f Yellow
       CmdAllowScripts
       CallMenu

      }
    '7' {
       #7) LISTS - Save as Templates
       $MenuSelect7 = "** Running option 7: LISTS - Save as Templates  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect7)" | Add-Content $logFile }
       Write-Host $MenuSelect7 -f Yellow 
       CmdListTemp 
       CallMenu

      }
    '8' {
       #8) LISTS - Download Templates
       $MenuSelect8 = "** Running option 8: LISTS - Download Templates  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect8)" | Add-Content $logFile }
       Write-Host $MenuSelect8 -f Yellow
       CmdDownTemp  
       CallMenu

      }
    '9' {
       #9) LISTS - Upload and Create
       $MenuSelect9 = "** Running option 9: LISTS - Upload and Create  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect9)" | Add-Content $logFile }
       Write-Host $MenuSelect9 -f Yellow
       CmdListCreate 
       CallMenu

      }
    '10' {
       #10) Assign Permissions
       $MenuSelect10 = "** Running option 10: Assign Permissions  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect10)" | Add-Content $logFile }
       Write-Host $MenuSelect10 -f Yellow
       CmdPerms
       CallMenu

      }
    '11' {
       #11) Export Navigation
       $MenuSelect11 = "** Running option 11: EXPORT Navigation  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect11)" | Add-Content $logFile }
       Write-Host $MenuSelect11 -f Yellow
       CmdExportNav
       CallMenu

      }
    '12' {
       #12) Import Navigation
       $MenuSelect12 = "** Running option 12: IMPORT Navigation  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect12)" | Add-Content $logFile }
       Write-Host $MenuSelect12 -f Yellow
       CmdImportNav
       CallMenu

      }
    '13' {
       #13) UTIL:Add Admin to ALL Sites
       $MenuSelect13 = "** Running option 13) UTIL:Add Admin to ALL Sites  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect13)" | Add-Content $logFile }
       Write-Host $MenuSelect13 -f Yellow
       CmdImportNav
       CallMenu

      }
    '14' {
       #14) UTIL:Remove Admin from ALL Sites
       $MenuSelect14 = "** Running option 14) UTIL:Remove Admin from ALL Sites  - $(Get-TimeStamp)"
       if ($log -eq $true) { "`n$($MenuSelect14)" | Add-Content $logFile }
       Write-Host $MenuSelect14 -f Yellow
       CmdImportNav
       CallMenu

      }

}
}Until($MenuResponse -eq 'x')


## Create Checklist 


