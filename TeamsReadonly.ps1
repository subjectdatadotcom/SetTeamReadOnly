<#
.SYNOPSIS
This script processes Teams information from a CSV file, retrieves associated SharePoint sites, sets their lock state to ReadOnly, and checks the lock state for verification.

.DESCRIPTION
The script ensures that the MicrosoftTeams, ExchangeOnlineManagement, and Microsoft.Online.SharePoint.PowerShell modules are installed and imported. It reads a list of Teams from a CSV file, connects to Microsoft Teams, Exchange Online, and SharePoint Online, retrieves the associated SharePoint sites for each Team, sets the site to ReadOnly, and checks the lock state, providing feedback on the operation.

.PARAMETER CSVFilePath
The path to the CSV file containing the Teams information. The file must include `GroupId` and `Source MailNickname` columns.

.NOTES
- The script requires administrative permissions for Microsoft Teams, Exchange Online, and SharePoint Online.
- Authentication is required for each service using valid credentials.

.AUTHOR
SubjectData

.EXAMPLE
.\TeamsReadonly.ps1
This runs the script in the current directory, processes the '131Teams.csv' file, and sets the associated SharePoint sites to ReadOnly.

.EXAMPLE
.\TeamsReadonly.ps1 -CSVFilePath "C:\Data\MyTeams.csv"
This runs the script and processes the Teams information from the specified CSV file path.
#>


$MicrosoftTeams = "MicrosoftTeams"

# Check if the module is already installed
if (-not(Get-Module -Name $MicrosoftTeams)) {
    # Install the module
    Install-Module -Name $MicrosoftTeams -Force
}

$MicrosoftOnline = "Microsoft.Online.Sharepoint.PowerShell"

# Check if the module is already installed
if (-not(Get-Module -Name $MicrosoftOnline)) {
    # Install the module
    Install-Module -Name $MicrosoftOnline -Force
}

$ExchangeOnlineManagement = "ExchangeOnlineManagement"
# Check if the module is already installed
if (-not(Get-Module -Name $ExchangeOnlineManagement)) {
    # Install the module
    Install-Module -Name $ExchangeOnlineManagement -Force
}

Import-Module -Name $ExchangeOnlineManagement

Import-Module $MicrosoftTeams -Force

Import-Module $MicrosoftOnline -Force


$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$XLloc = "$myDir\"

try
{
    $TeamsList = import-csv ($XLloc+"131Teams.csv").ToString()
}
catch
{
    Write-Host "No CSV file to read" -BackgroundColor Black -ForegroundColor Red
    exit
}

###########################
###   Connect MSTeams   ###
###########################

Connect-MicrosoftTeams

Connect-ExchangeOnline

Connect-SPOService -url "https://coxautoinc-admin.sharepoint.com"

if($TeamsList.Count -gt 0)
{
    foreach ($Team in $TeamsList)
    {   
        try
        {
            Write-Host "$($Team.'Source MailNickname')"

            $currentTeam = Get-Team -GroupId $Team.'GroupId'

            #$teamGroupId = $varNewTeam.GroupId
        
            #Run a Get-UnifiedGroup off of the Teams DisplayNames

            # $associatedTeam = Get-UnifiedGroup | Where-Object { $_.SharePointSiteUrl -eq "https://tradercorporation.sharepoint.com/sites/CMS-CIBC-2023RFPTeam"}
            # $associatedTeam.DisplayName

            $sites = (Get-UnifiedGroup  $currentTeam.DisplayName).SharePointSiteUrl
            
            #Loop through the Sharepoint Sites Displaying Title & Lockstate

            #foreach ($site in $sites) { Get-SPOSite  $site | select Title, Lockstate}

            #Loop through the Sharepoint Sites setting the Lockstate to ReadOnly

            foreach ($site in $sites) 
            { 
                Write-Host "Setting site readonly $($site)" -ForegroundColor Yellow
                Set-SPOSite $site -Lockstate ReadOnly
            }

            #Loop through the Sharepoint Sites Displaying Titale & Lockstate after changes

            #foreach ($site in $sites) { Get-SPOSite  $site | select Title, Lockstate}

        }
        catch
        {
            Write-Host "Exception occured " $Team.'Source MailNickname' -BackgroundColor Black -ForegroundColor Red
        }
    }
}



